unit MSJet;

{$WARN UNSAFE_TYPE OFF}
{$WARN UNSAFE_CODE OFF}
{$WARN UNSAFE_CAST OFF}

interface

{$ALIGN 4}

uses Winapi.Windows, System.Classes, System.SysUtils, System.Variants, System.Generics.Collections,
  MSJet.Consts, MSJet.API;

type
  TMSJetObject  = class;
  TMSJetSession = class;
  TMSJetTable   = class;

  {$region 'Base classes'}

  TMSJetItemState = (itsUnmodified, itsNew, itsModified, itsDelete, itsDeleted);

  TMSJetObjectClass = class of TMSJetObject;

  TMSJetObjectBaseList = class;

  TMSJetObject = class(TPersistent)
  private
    FName: AnsiString;
    FItemState: TMSJetItemState;
    FList: TMSJetObjectBaseList;
    function  GetName: string;
    function  GetIndex: Integer;
    procedure SetIndex(const Value: Integer);
  protected
    property  List: TMSJetObjectBaseList read FList;
  public
    destructor Destroy; override;
    property  Index: Integer read GetIndex write SetIndex;
    property  ItemState: TMSJetItemState read FItemState write FItemState;
  published
    property  Name: string read GetName;
  end;

  TMSJetObjectBaseList = class(TPersistent)
  private
    FList: TList;
  protected
    function  DoRemove(Item: TMSJetObject): Integer; virtual; abstract;
  end;

  TMSJetObjectList<TItem: TMSJetObject> = class(TMSJetObjectBaseList)
  private
    FItems: TDictionary<AnsiString, TItem>;
    function  GetCount: Integer; inline;
    function  Get(Index: Integer): TItem; inline;
    procedure Put(Index: Integer; const Value: TItem); inline;
  protected
    function  CreateItem(const Name: string): TItem; virtual; abstract;
    function  DoRemove(Item: TMSJetObject): Integer; override;
  public
    constructor Create;
    destructor Destroy; override;
    function  Add(const Name: string): TItem;
    function  ByName(const Name: string): TItem; inline;
    procedure Clear;
    procedure Delete(Index: Integer);
    function  Find(const Name: string): TItem; inline;
    function  FindOrAdd(const Name: string): TItem;
    function  Exists(const Name: string): Boolean;
    function  IndexOf(const Name: string): Integer;
    function  Insert(Index: Integer; const Name: string): TItem;
    function  Remove(Item: TItem): Integer; inline;
    property  Count: Integer read GetCount;
    property  Items[Index: Integer]: TItem read Get write Put; default;
  end;

  TMSJetObjectInfo = class(TMSJetObject)
  private
    FContainer: string;
    FFlags: ULONG;
    FOptions: ULONG;
  published
    property  Container: string read FContainer write FContainer;
    property  Options: ULONG read FOptions write FOptions;
    property  Flags: ULONG read FFlags write FFlags;
  end;

  TMSJetObjectInfoList = class(TMSJetObjectList<TMSJetObjectInfo>)
  protected
    function  CreateItem(const Name: string): TMSJetObjectInfo; override;
  end;

  TMSJetMetaObject = class(TMSJetObject)
  private
    FHandle: ULONG;
    FOriginalName: string;
    function  GetRealName: string; inline;
  protected
    procedure SetHandle(const Value: JET_COLUMNID); virtual;
  protected
    procedure DoDbDelete(const Name: AnsiString); virtual; abstract;
    procedure DoDbCreate(const Name: AnsiString); virtual; abstract;
    procedure DoDbRename(const OldName, NewName: AnsiString); virtual; abstract;
    procedure DoDbUpdate(const Name: AnsiString); virtual; abstract;
    procedure DoDbReload(const Name: AnsiString); virtual; abstract;
  public
    procedure DbDelete;
    procedure DbCreate;
    procedure DbRename(const NewName: string);
    procedure DbUpdate;
    procedure DbReload;
    property  Handle: ULONG read FHandle write SetHandle;
    property  OriginalName: string read FOriginalName write FOriginalName;
    property  RealName: string read GetRealName;
  end;

  TMSJetMetaObjectList<TItem: TMSJetMetaObject> = class(TMSJetObjectList<TItem>)
  public
    procedure Reload; virtual; abstract;
    procedure Apply;
  end;

  TMSJetTableItem = class(TMSJetMetaObject)
  private
    FTable: TMSJetTable;
  protected
  public
    constructor Create(Table: TMSJetTable); virtual;
    property  Table: TMSJetTable read FTable;
  end;

  TMSJetTableItemList<TItem: TMSJetTableItem> = class(TMSJetMetaObjectList<TItem>)
  private
    FTable: TMSJetTable;
  public
    constructor Create(Table: TMSJetTable);
    property  Table: TMSJetTable read FTable;
  end;

  {$endregion}

  {$region 'Jet'}

  TMSJetInstance = class(TPersistent)
  private
    FName: AnsiString;
    FHandle: JET_INSTANCE;
    FPUCW: Word;
    FMaxVerPages: Cardinal;
    function  GetName: string;
  protected
    procedure BeginCall;
    procedure EndCall;
  public
    constructor Create(const Name: string);
    destructor Destroy; override;
    procedure Open;
    procedure Close;
    function  CreateSession: TMSJetSession;
    property  Name: string read GetName;
    property  MaxVerPages: Cardinal read FMaxVerPages write FMaxVerPages;
    property  Handle: JET_INSTANCE read FHandle;
  end;

  TMSJetSession = class(TPersistent)
  private
    FInstance: TMSJetInstance;
    FHandle: JET_SESID;
  public
    constructor Create(Instance: TMSJetInstance); overload;
    constructor Create(Handle: JET_SESID); overload;
    destructor Destroy; override;
    procedure BeginTransaction;
    procedure Commit;
    procedure Rollback;
    property  Instance: TMSJetInstance read FInstance;
    property  Handle: JET_SESID read FHandle;
  end;

  TMSJetDatabase = class(TPersistent)
  private
    FSession: TMSJetSession;
    FHandle: JET_DBID;
    FFileName: AnsiString;
    FOwnHandle: Boolean;
    function  GetFileName: string;
    procedure SetFileName(const Value: string);
    procedure SetHandle(const Value: JET_DBID);
  public
    constructor Create(Session: TMSJetSession);
    destructor Destroy; override;
    function  Select(const TableName, IndexName: string; const Keys: array of Variant): TMSJetTable;
    procedure CreateDatabase;
    procedure Open;
    procedure Close;
    procedure ReadTableList(Items: TMSJetObjectInfoList);
    property  FileName: string read GetFileName write SetFileName;
    property  Session: TMSJetSession read FSession;
    property  Handle: JET_DBID read FHandle write SetHandle;
  end;

  {$endregion}

  {$region 'Columns'}

  TMSJetColumnType = (
    jetBit = JET_coltypBit,
    jetUByte = JET_coltypUnsignedByte,
    jetShort = JET_coltypShort,
    jetLong = JET_coltypLong,
    jetCurrency = JET_coltypCurrency,
    jetSingle = JET_coltypIEEESingle,
    jetDouble = JET_coltypIEEEDouble,
    jetDateTime = JET_coltypDateTime,
    jetBinary = JET_coltypBinary,
    jetText = JET_coltypText,
    jetLongBinary = JET_coltypLongBinary,
    jetLongText = JET_coltypLongText,
    jetSLV = JET_coltypSLV,
    jetULong = JET_coltypUnsignedLong,
    jetLongLong = JET_coltypLongLong,
    jetGUID = JET_coltypGUID,
    jetUShort = JET_coltypUnsignedShort
  );

  TMSJetColumnOption = (
    jetColumnFixed,
    jetColumnTagged,
    jetColumnNotNULL,
    jetColumnVersion,
    jetColumnAutoincrement,
    jetColumnUpdatable,
    jetColumnTTKey,
    jetColumnTTDescending,
    jetColumnMultiValued,
    jetColumnEscrowUpdate,
    jetColumnUnversioned,
    jetDeleteOnZero,
    jetColumnMaybeNull,
    jetColumnFinalize,
    jetColumnUserDefinedDefault
  );
  TMSJetColumnOptions = set of TMSJetColumnOption;

  EMSJetColumnTypeError = class(EMSJetError)
  public
    constructor Create(const ColumnName: string; DataType: TMSJetColumnType);
  end;

  TMSJetColumn = class(TMSJetTableItem)
  private
    FDataType: TMSJetColumnType;
    FSize: Cardinal;
    FOptions: TMSJetColumnOptions;
    FDisplayName: string;
    function  GetAsBoolean: Boolean;
    function  GetAsCurrency: Currency;
    function  GetAsDateTime: TDateTime;
    function  GetAsFloat: Double;
    function  GetAsInt64: Int64;
    function  GetAsInteger: Integer;
    function  GetAsSingle: Single;
    function  GetAsString: string;
    function  GetIsNull: Boolean;
    function  GetValue: Variant;
    procedure SetAsString(const Value: string);
    procedure SetAsInteger(const Value: Integer);
    procedure SetAsBoolean(const Value: Boolean);
    procedure SetAsCurrency(const Value: Currency);
    procedure SetAsDateTime(const Value: TDateTime);
    procedure SetAsFloat(const Value: Double);
    procedure SetAsInt64(const Value: Int64);
    procedure SetAsSingle(const Value: Single);
    procedure SetIsNull(const Value: Boolean);
    procedure SetValue(const Value: Variant);
    function  GetAsRaw: RawByteString;
    procedure SetAsRaw(const Value: RawByteString);
    procedure SetDisplayName(const Value: string);
  protected
    procedure SetHandle(const Value: JET_COLUMNID); override;
  protected
    procedure DoDbDelete(const Name: AnsiString); override;
    procedure DoDbCreate(const Name: AnsiString); override;
    procedure DoDbRename(const OldName, NewName: AnsiString); override;
    procedure DoDbUpdate(const Name: AnsiString); override;
    procedure DoDbReload(const Name: AnsiString); override;
  public
    constructor Create(Table: TMSJetTable); override;
    procedure AssignTo(Dest: TPersistent); override;
    function  Retrieve(var Value): ULONG;
    procedure Update(const P: Pointer; Size: Cardinal);
    property  AsBoolean: Boolean read GetAsBoolean write SetAsBoolean;
    property  AsCurrency: Currency read GetAsCurrency write SetAsCurrency;
    property  AsDateTime: TDateTime read GetAsDateTime write SetAsDateTime;
    property  AsFloat: Double read GetAsFloat write SetAsFloat;
    property  AsInt64: Int64 read GetAsInt64 write SetAsInt64;
    property  AsInteger: Integer read GetAsInteger write SetAsInteger;
    property  AsSingle: Single read GetAsSingle write SetAsSingle;
    property  AsString: string read GetAsString write SetAsString;
    property  AsRaw: RawByteString read GetAsRaw write SetAsRaw;
    property  IsNull: Boolean read GetIsNull write SetIsNull;
    property  Value: Variant read GetValue write SetValue;
  published
    property  DataType: TMSJetColumnType read FDataType write FDataType;
    property  DisplayName: string read FDisplayName write SetDisplayName;
    property  Options: TMSJetColumnOptions read FOptions write FOptions;
    property  Size: Cardinal read FSize write FSize;
  end;

  TMSJetColumnList = class(TMSJetTableItemList<TMSJetColumn>)
  protected
    function  CreateItem(const Name: string): TMSJetColumn; override;
  public
    function  Add(const Name: string; DataType: TMSJetColumnType; Size: Cardinal = 0; Options: TMSJetColumnOptions = []): TMSJetColumn; overload;
    procedure Reload; override;
  end;

  {$endregion}

  {$region 'Index'}

  TMSJetIndexOption = (
    jetIndexUnique,
    jetIndexPrimary,
    jetIndexDisallowNull,
    jetIndexIgnoreNull,
    jetIndexIgnoreAnyNull,
    jetIndexIgnoreFirstNull,
    jetIndexLazyFlush,
    jetIndexEmpty,
    jetIndexUnversioned,
    jetIndexSortNullsHigh,
    jetIndexUnicode,
    jetIndexTuples,
    jetIndexTupleLimits,
    jetIndexCrossProduct,
    jetIndexKeyMost,
    jetIndexDisallowTruncation,
    jetIndexNestedTable
  );
  TMSJetIndexOptions = set of TMSJetIndexOption;

  TMSJetIndex = class(TMSJetTableItem)
  private
    FFields: string;
    FKey: AnsiString;
    FDensity: Cardinal;
    FOptions: TMSJetIndexOptions;
    FColumns: TList;
    procedure SetFields(const Value: string);
    function  GetColumn(const Index: Integer): TMSJetColumn; inline;
    function  GetColumnCount: Integer; inline;
  protected
    procedure DoDbDelete(const Name: AnsiString); override;
    procedure DoDbCreate(const Name: AnsiString); override;
    procedure DoDbRename(const OldName, NewName: AnsiString); override;
    procedure DoDbUpdate(const Name: AnsiString); override;
    procedure DoDbReload(const Name: AnsiString); override;
  public
    constructor Create(Table: TMSJetTable); override;
    destructor Destroy; override;
    procedure AssignTo(Dest: TPersistent); override;
    property  Key: AnsiString read FKey;
    property  ColumnCount: Integer read GetColumnCount;
    property  Columns[const Index: Integer]: TMSJetColumn  read GetColumn;
  published
    property  Fields: string read FFields write SetFields;
    property  Options: TMSJetIndexOptions read FOptions write FOptions;
    property  Density: Cardinal read FDensity write FDensity;
  end;

  TMSJetIndexList = class(TMSJetTableItemList<TMSJetIndex>)
  protected
    function  CreateItem(const Name: string): TMSJetIndex; override;
  public
    function  Add(const Name, Columns: string; Options: TMSJetIndexOptions = []): TMSJetIndex; overload;
    procedure Reload; override;
  end;

  {$endregion}

  {$region 'Table'}

  TMSJetMove = (
    jetMoveFirst = JET_MoveFirst,
    jetMovePreious = JET_MovePrevious,
    jetMoveNext = JET_MoveNext,
    jetMoeLast = JET_MoveLast
  );

  TMSJetKeyOption = (
    jetNewKey,
    jetStrLimit,
    jetSubStrLimit,
    jetNormalizedKey,
    jetKeyDataZeroLength,
    jetFullColumnStartLimit,
    jetFullColumnEndLimit,
    jetPartialColumnStartLimit,
    jetPartialColumnEndLimit
  );
  TMSJetKeyOptions = set of TMSJetKeyOption;

  TMSJetSeekOption = (
    jetSeekEQ = JET_bitSeekEQ,
    jetSeekLT = JET_bitSeekLT,
    jetSeekLE = JET_bitSeekLE,
    jetSeekGE = JET_bitSeekGE,
    jetSeekGT = JET_bitSeekGT
  );
  TMSJetSeekOptions = set of TMSJetSeekOption;

  JetBookmark = interface
  ['{03580CAE-6536-4F51-B440-F5BA939277C9}']
    function Go: Boolean;
  end;

  TJetBookmark = class(TInterfacedObject, JetBookmark)
  private
    FData: Pointer;
    FSize: ULONG;
    FSecondaryData: Pointer;
    FSecondarySize: ULONG;
    FTable: TMSJetTable;
  protected
    function Go: Boolean;
  public
    constructor Create(Table: TMSJetTable; const Data: Pointer; const Size: ULONG; const SecondaryData: Pointer; const SecondarySize: ULONG);
    destructor Destroy; override;
  end;

  TMSJetJoin = class(TPersistent)
  private
    FTable: TMSJetTable;
    FParentKey: string;
    FParent: TMSJetTable;
    FIsSync: Boolean;
    FColumns: TArray<string>;
    function  GetValue(const Name: string): Variant;
  public
    constructor Create(Parent: TMSJetTable; const TableName, IndexName, ParentKey: string);
    destructor Destroy; override;
    function  Sync(IsSync: Boolean): Boolean;
    property  Table: TMSJetTable read FTable;
    property  Parent: TMSJetTable read FParent;
    property  ParentKey: string read FParentKey;
    property  IsSync: Boolean read FIsSync;
    property  Columns: TArray<string> read FColumns write FColumns;
    property  Value[const Name: string]: Variant read GetValue; default;
  end;

  TMSJetTable = class(TMSJetMetaObject)
  private
    FColumns: TMSJetColumnList;
    FDatabase: TMSJetDatabase;
    FPages: Cardinal;
    FDensity: Cardinal;
    FIndexes: TMSJetIndexList;
    FIndexName: AnsiString;
    FKey: Pointer;
    FKeySize: ULONG;
    FSeekKeys: array of Variant;
    FSeekOption: TMSJetSeekOption;
    FSeekSetRange: Boolean;
    FIsInsert: Boolean;
    FJoins: TDictionary<string, TMSJetJoin>;
    procedure SetColumns(const Value: TMSJetColumnList);
    procedure SetIndexes(const Value: TMSJetIndexList);
    function  GetIndexName: string;
    procedure SetIndexName(const Value: string);
    function  GetColumn(const ColumnName: string): TMSJetColumn;
    function  GetLink(const ParentKey: string): TMSJetJoin;
    function  GetJoins: TArray<TMSJetJoin>;
    function  GetJoinCount: Integer;
  protected
    function  GetColumnText(id: JET_COLUMNID): string;
    procedure MakeKey(Column: TMSJetColumn; const Value: Variant; Options: TMSJetKeyOptions);
    function  CreatePrimaryBookmark: JetBookmark;
    function  CreateSecondaryBookmark: JetBookmark;
    procedure SyncCursor(HasCursor: Boolean);
  protected
    procedure DoDbDelete(const Name: AnsiString); override;
    procedure DoDbCreate(const Name: AnsiString); override;
    procedure DoDbRename(const OldName, NewName: AnsiString); override;
    procedure DoDbUpdate(const Name: AnsiString); override;
    procedure DoDbReload(const Name: AnsiString); override;
    property  IsInsert: Boolean read FIsInsert;
  public
    constructor Create(Database: TMSJetDatabase; const Name: string); overload;
    constructor Create(Database: TMSJetDatabase; Handle: JET_TABLEID); overload;
    destructor Destroy; override;
    procedure Open;
    procedure Close;
    function  First: Boolean;
    function  Last: Boolean;
    function  Next: Boolean;
    function  Previous: Boolean;
    procedure Insert;
    procedure Edit;
    procedure Cancel;
    procedure Post;
    procedure Delete;
    function  AddColumn(const Name: string; DataType: TMSJetColumnType; Size: ULONG = 0; Options: TMSJetColumnOptions = []): TMSJetColumn;
    function  Seek(const Keys: array of Variant; Option: TMSJetSeekOption = jetSeekEQ; SetRange: Boolean = False): Boolean;
    procedure ClearRange;
    function  RetrieveKey: Boolean;
    function  GoToKey: Boolean;
    function  CreateBookmark: JetBookmark;
    function  GotoBookmark(const Bookmark: JetBookmark): Boolean;
    function  Join(const TableName, IndexName, ParentKey: string): TMSJetJoin;
    property  IndexName: string read GetIndexName write SetIndexName;
    property  Cols[const ColumnName: string]: TMSJetColumn read GetColumn; default;
    property  Link[const ParentKey: string]: TMSJetJoin read GetLink;
    property  Joins: TArray<TMSJetJoin> read GetJoins;
    property  JoinCount: Integer read GetJoinCount;
  published
    property  Database: TMSJetDatabase read FDatabase;
    property  Columns: TMSJetColumnList read FColumns write SetColumns;
    property  Indexes: TMSJetIndexList read FIndexes write SetIndexes;
    property  Pages: Cardinal read FPages write FPages;
    property  Density: Cardinal read FDensity write FDensity;
  end;

  TMSJetTableList = class(TMSJetMetaObjectList<TMSJetTable>)
  private
    FDatabase: TMSJetDatabase;
  protected
    function  CreateItem(const Name: string): TMSJetTable; override;
  public
    constructor Create(Database: TMSJetDatabase);
    property  Database: TMSJetDatabase read FDatabase;
  end;

  {$endregion}

  {$region 'Utils'}

const
  MSJetTypeNames: array[TMSJetColumnType] of string = (
    'Bit',
    'UByte',
    'Short',
    'Long',
    'Currency',
    'Single',
    'Double',
    'DateTime',
    'Binary',
    'Text',
    'LongBinary',
    'LongText',
    'SLV',
    'ULong',
    'LongLong',
    'GUID',
    'UShort'
  );

  MSJetColumnOptionNames: array[TMSJetColumnOption] of string = (
    'Fixed',
    'Tagged',
    'NotNULL',
    'Version',
    'Autoincrement',
    'Updatable',
    'TTKey',
    'TTDescending',
    'MultiValued',
    'EscrowUpdate',
    'Unversioned',
    'DeleteOnZero',
    'MaybeNull',
    'Finalize',
    'UserDefinedDefault'
  );

  MSJetIndexOptionNames: array[TMSJetIndexOption] of string = (
    'Unique',
    'Primary',
    'DisallowNull',
    'IgnoreNull',
    'IgnoreAnyNull',
    'IgnoreFirstNull',
    'LazyFlush',
    'Empty',
    'Unversioned',
    'SortNullsHigh',
    'Unicode',
    'Tuples',
    'TupleLimits',
    'CrossProduct',
    'KeyMost',
    'DisallowTruncation',
    'NestedTable'
  );

  function  SetToInt(const Value): Cardinal;
  procedure IntToSet(Value: Cardinal; out Flags);

  {$endregion}


implementation

uses StrUtils;

function SetToInt(const Value): Cardinal;
begin
  Result := PDWord(@Value)^;
end;

procedure IntToSet(Value: Cardinal; out Flags);
begin
  PDWord(@Flags)^ := Value;
end;

{ EMSJetColumnTypeError }

constructor EMSJetColumnTypeError.Create(const ColumnName: string; DataType: TMSJetColumnType);
begin
  inherited Create('Could not access column "' + ColumnName + '" as "' + MSJetTypeNames[DataType] + '"');
end;

{ TMSJetObject }

destructor TMSJetObject.Destroy;
begin
  if FList <> nil then
    FList.DoRemove(Self);
  inherited Destroy;
end;

function TMSJetObject.GetIndex: Integer;
begin
  Result := List.FList.IndexOf(Self);
end;

function  TMSJetObject.GetName: string;
begin
  Result := string(FName);
end;

procedure TMSJetObject.SetIndex(const Value: Integer);
begin
  List.FList.Move(Index, Value);
end;

{ TMSJetObjectList<TItem> }

constructor TMSJetObjectList<TItem>.Create;
begin
  inherited Create;
  FList := TList.Create;
  FItems := TDictionary<AnsiString, TItem>.Create;
end;

destructor TMSJetObjectList<TItem>.Destroy;
begin
  Clear;
  FreeAndNil(FList);
  FreeAndNil(FItems);
  inherited;
end;

procedure TMSJetObjectList<TItem>.Clear;
var
  I: Integer;
begin
  for I := 0 to FList.Count - 1 do
  begin
    Items[I].Flist := nil;
    Items[I].Free;
  end;
  FList.Clear;
  FItems.Clear;
end;

function TMSJetObjectList<TItem>.IndexOf(const Name: string): Integer;
var
  I: Integer;
begin
  for I := 0 to Count - 1 do
    if Items[I].Name = Name then
    begin
      Result := I;
      Exit;
    end;
  Result := -1;
end;

function TMSJetObjectList<TItem>.Find(const Name: string): TItem;
begin
  if not FItems.TryGetValue(AnsiString(Name), Result) then
    Result := nil;
end;

function TMSJetObjectList<TItem>.Exists(const Name: string): Boolean;
var
  Item: TItem;
begin
  Result := FItems.TryGetValue(AnsiString(Name), Item);
end;

function TMSJetObjectList<TItem>.ByName(const Name: string): TItem;
begin
  Result := Find(Name);
  if Result = nil then
    raise EMSJetError.Create('Item "' + Name + '" not found');
end;

function TMSJetObjectList<TItem>.FindOrAdd(const Name: string): TItem;
begin
  Result := Find(Name);
  if Result = nil then
    Result := Add(Name);
end;

function TMSJetObjectList<TItem>.Add(const Name: string): TItem;
begin
  Result := CreateItem(Name);
  Result.ItemState := itsNew;
  Result.FList := Self;
  FList.Add(Pointer(Result));
  FItems.Add(AnsiString(Name), Result);
end;

function TMSJetObjectList<TItem>.Insert(Index: Integer; const Name: string): TItem;
begin
  Result := CreateItem(Name);
  Result.ItemState := itsNew;
  Result.FList := Self;
  FList.Insert(Index, Pointer(Result));
  FItems.Add(AnsiString(Name), Result);
end;

procedure TMSJetObjectList<TItem>.Delete(Index: Integer);
begin
  FItems.Remove(AnsiString(Items[Index].Name));
  Items[Index].FList := nil;
  Items[Index].Free;
  FList.Delete(Index);
end;

function TMSJetObjectList<TItem>.DoRemove(Item: TMSJetObject): Integer;
begin
  Result := FList.IndexOf(Pointer(Item));
  if Result >= 0 then
  begin
    FItems.Remove(AnsiString(Items[Result].Name));
    FList.Delete(Result);
  end;
end;

function TMSJetObjectList<TItem>.Remove(Item: TItem): Integer;
begin
  Result := DoRemove(Item);
end;

function TMSJetObjectList<TItem>.GetCount: Integer;
begin
  Result := FList.Count;
end;

function TMSJetObjectList<TItem>.Get(Index: Integer): TItem;
begin
  Result := FList[Index];
end;

procedure TMSJetObjectList<TItem>.Put(Index: Integer; const Value: TItem);
begin
  FItems.Remove(AnsiString(Items[Index].Name));
  FList[Index] := Pointer(Value);
  FItems.Add(AnsiString(Value.Name), Value);
end;

{ TMSJetObjectInfoList }

function TMSJetObjectInfoList.CreateItem(const Name: string): TMSJetObjectInfo;
begin
  Result := TMSJetObjectInfo.Create;
  Result.FName := AnsiString(Name);
end;

{ TMSJetInstance }

constructor TMSJetInstance.Create(const Name: string);
begin
  inherited Create;
  FName := AnsiString(Name);

  FPUCW := Get8087CW;
  FHandle := 0;
  FMaxVerPages := 16 * 1024;
end;

destructor TMSJetInstance.Destroy;
begin
  Close;
  inherited;
end;

procedure TMSJetInstance.BeginCall;
begin
  //Some ESE calls requires specific FPU flags
  Set8087CW($027F);
end;

procedure TMSJetInstance.EndCall;
begin
  Set8087CW(FPUCW);
end;

procedure TMSJetInstance.Open;
begin
  BeginCall;
  try
    JetError(JetCreateInstance(FHandle, PAnsiChar(FName)));

    JetError(JetSetSystemParameter(FHandle, 0, JET_paramMaxVerPages, MaxVerPages, nil));

    JetError(JetInit(FHandle));
  finally
    EndCall;
  end;
end;

procedure TMSJetInstance.Close;
begin
  if FHandle <> 0 then
  begin
    JetError(JetTerm(FHandle));
    FHandle := 0;
  end;
end;

function TMSJetInstance.CreateSession: TMSJetSession;
begin
  Result := TMSJetSession.Create(Self);
end;

function TMSJetInstance.GetName: string;
begin
  Result := string(FName);
end;

{ TMSJetSession }

constructor TMSJetSession.Create(Instance: TMSJetInstance);
begin
  inherited Create;
  FInstance := Instance;

  Instance.BeginCall;
  try
    JetError(JetBeginSession(Instance.Handle, FHandle, nil, nil));
  finally
    Instance.EndCall;
  end;
end;

constructor TMSJetSession.Create(Handle: JET_SESID);
begin
  inherited Create;
  FHandle := Handle;
end;

destructor TMSJetSession.Destroy;
begin
  if Instance <> nil then
  begin
    Instance.BeginCall;
    try
      JetError(JetEndSession(FHandle, 0));
    finally
      Instance.EndCall;
    end;
  end;

  inherited;
end;

procedure TMSJetSession.BeginTransaction;
begin
  JetError(JetBeginTransaction(Handle));
end;

procedure TMSJetSession.Commit;
begin
  JetError(JetCommitTransaction(Handle, 0));
end;

procedure TMSJetSession.Rollback;
begin
  JetError(JetRollback(Handle, 0));
end;

{ TMSJetDatabase }

constructor TMSJetDatabase.Create(Session: TMSJetSession);
begin
  inherited Create;
  FSession := Session;
  FHandle := 0;
end;

destructor TMSJetDatabase.Destroy;
begin
  Close;
  inherited;
end;

procedure TMSJetDatabase.CreateDatabase;
begin
  Close;
  JetError(JetCreateDatabase(Session.Handle, PAnsiChar(FFileName), nil, FHandle, 0));
end;

procedure TMSJetDatabase.Open;
begin
  Close;
  JetError(JetAttachDatabase(Session.Handle, PAnsiChar(FFileName), 0));
  JetError(JetOpenDatabase(Session.Handle, PAnsiChar(FFileName), nil, FHandle, 0));
  FOwnHandle := True;
end;

procedure TMSJetDatabase.Close;
begin
  if FOwnHandle and (FHandle <> 0) then
  begin
    JetError(JetCloseDatabase(Session.Handle, FHandle, 0));
    JetError(JetDetachDatabase(Session.Handle, PAnsiChar(FFileName)));
    FHandle := 0;
  end;
end;

procedure TMSJetDatabase.SetHandle(const Value: JET_DBID);
begin
  Close;
  FHandle := Value;
  FOwnHandle := False;
end;

function TMSJetDatabase.GetFileName: string;
begin
  Result := string(FFileName);
end;

function TMSJetDatabase.Select(const TableName, IndexName: string; const Keys: array of Variant): TMSJetTable;
begin
  Result := TMSJetTable.Create(Self, TableName);
  try
    Result.IndexName := IndexName;
    Result.Open;
    if not Result.Seek(Keys, jetSeekEQ, True) then
      FreeAndNil(Result);
  except
    Result.Free;
    raise;
  end;
end;

procedure TMSJetDatabase.SetFileName(const Value: string);
begin
  if FFileName <> AnsiString(Value) then
  begin
    Close;
    FFileName := AnsiString(Value);
  end;
end;

procedure TMSJetDatabase.ReadTableList(Items: TMSJetObjectInfoList);
var
  List: JET_OBJECTLIST;
  Table: TMSJetTable;
  Item: TMSJetObjectInfo;
  BufSize: ULONG;
begin
  ZeroMemory(@List, SizeOf(List));
  List.cbStruct := SizeOf(List);
  JetError(JetGetObjectInfo(Session.Handle, Handle, JET_objtypTable, nil, nil, @List, SizeOf(List), 1));
  Table := TMSJetTable.Create(Self, List.tableid);
  try
    if Table.Last then
    repeat
      Item := Items.Add(Table.GetColumnText(List.columnidobjectname));

      Item.Container := Table.GetColumnText(List.columnidcontainername);

      JetError(JetRetrieveColumn(Session.Handle, Table.Handle, List.columnidgrbit, @Item.FOptions, SizeOf(Item.FOptions), BufSize, 0, nil));
      JetError(JetRetrieveColumn(Session.Handle, Table.Handle, List.columnidflags, @Item.FFlags, SizeOf(Item.FFlags), BufSize, 0, nil));
    until not Table.Previous;
  finally
    Table.Free;
  end;
end;

{ TMSJetMetaObject }

procedure TMSJetMetaObject.DbCreate;
begin
  DoDbCreate(FName);
  ItemState := itsUnmodified;
end;

procedure TMSJetMetaObject.DbDelete;
begin
  DoDbDelete(AnsiString(RealName));
  ItemState := itsDeleted;
end;

procedure TMSJetMetaObject.DbRename(const NewName: string);
begin
  DoDbRename(AnsiString(RealName), AnsiString(NewName));
  OriginalName := '';
  FName := AnsiString(NewName);
end;

procedure TMSJetMetaObject.DbUpdate;
begin
  if Handle = 0 then
  begin
    DbCreate;
  end
  else
  begin
    if RealName <> Name then
    begin
      DoDbRename(AnsiString(RealName), AnsiString(Name));
      OriginalName := '';
    end;
    DoDbUpdate(AnsiString(Name));
    ItemState := itsUnmodified;
  end;
end;

procedure TMSJetMetaObject.DbReload;
begin
  DoDbReload(AnsiString(RealName));
  FName := AnsiString(RealName);
  OriginalName := '';
  ItemState := itsUnmodified;
end;

function TMSJetMetaObject.GetRealName: string;
begin
  if OriginalName <> '' then
    Result := OriginalName
  else
    Result := Name;
end;

procedure TMSJetMetaObject.SetHandle(const Value: JET_COLUMNID);
begin
  FHandle := Value;
end;

{ TMSJetMetaObjectList<TItem> }

procedure TMSJetMetaObjectList<TItem>.Apply;
var
  I: Integer;
  Item: TMSJetMetaObject;
begin
  I := 0;
  while I < Count do
  begin
    Item := Items[I];
    case Item.ItemState of
      itsNew:
        begin
          Item.DbCreate;
          Inc(I);
        end;
      itsDelete:
        begin
          Item.DbDelete;
          Remove(Item);
        end;
      itsModified:
        begin
          Item.DbUpdate;
          Inc(I);
        end;
      else
        begin
          Inc(I);
        end;
    end;

  end;

end;

{ TMSJetTableItem }

constructor TMSJetTableItem.Create(Table: TMSJetTable);
begin
  inherited Create;
  FTable := Table;
end;
{ TMSJetTableList<TItem: TMSJetTableItem> }

constructor TMSJetTableItemList<TItem>.Create(Table: TMSJetTable);
begin
  inherited Create;
  FTable := Table;
end;

{ TMSJetColumn }

constructor TMSJetColumn.Create(Table: TMSJetTable);
begin
  inherited Create(Table);
  FDataType := jetLong;
end;

procedure TMSJetColumn.DoDbCreate(const Name: AnsiString);
var
  Def: JET_COLUMNDEF;
begin
  ZeroMemory(@Def, SizeOf(Def));
  Def.cbStruct := SizeOf(Def);
  Def.coltyp := ULONG(DataType);
  Def.cbMax := Size;
  Def.grbit := SetToInt(Options);
  Def.cp := 1200; // Unicode

  JetError(JetAddColumn(Table.Database.Session.Handle, Table.Handle, PAnsiChar(Name), @Def, nil, 0, FHandle));
end;

procedure TMSJetColumn.DoDbDelete(const Name: AnsiString);
begin
  JetError(JetDeleteColumn(Table.Database.Session.Handle, Table.Handle, PAnsiChar(Name)));
  FHandle := 0;
end;

procedure TMSJetColumn.DoDbRename(const OldName, NewName: AnsiString);
begin
  JetError(JetRenameColumn(Table.Database.Session.Handle, Table.Handle, PAnsiChar(OldName), PAnsiChar(NewName), 0));
end;

procedure TMSJetColumn.DoDbUpdate(const Name: AnsiString);
begin

end;

procedure TMSJetColumn.DoDbReload(const Name: AnsiString);
var
  Info: JET_COLUMNBASE;
begin
  ZeroMemory(@Info, SizeOf(Info));
  Info.cbStruct := SizeOf(Info);
  JetError(JetGetTableColumnInfo(Table.Database.Session.Handle, Table.Handle, @FHandle, @Info, SizeOf(Info), 8));

  FName := AnsiString(Info.szBaseColumnName);
  FDataType := TMSJetColumnType(Info.coltyp);
  FSize := Info.cbMax;
  IntToSet(Info.grbit, FOptions);
end;

procedure TMSJetColumn.AssignTo(Dest: TPersistent);
begin
  if Dest is TMSJetColumn then
  begin
    TMSJetColumn(Dest).FDataType := DataType;
    TMSJetColumn(Dest).FName := FName;
    TMSJetColumn(Dest).FSize := Size;
    TMSJetColumn(Dest).FOptions := Options;
  end
  else
    inherited;
end;

function TMSJetColumn.Retrieve(var Value): ULONG;
var
  S: string;
begin
  case DataType of
  jetBinary, jetLongBinary:
    begin
      JetError(JetRetrieveColumn(Table.Database.Session.Handle, Table.Handle, Handle, nil, 0, Result, 0, nil));
      if Result > 0 then
      begin
        ReallocMem(Pointer(Value), Result);
        JetError(JetRetrieveColumn(Table.Database.Session.Handle, Table.Handle, Handle, Pointer(Value), Result, Result, 0, nil));
      end;
    end;
  jetText, jetLongText:
    begin
      JetError(JetRetrieveColumn(Table.Database.Session.Handle, Table.Handle, Handle, nil, 0, Result, 0, nil));
      if Result > 0 then
      begin
        SetLength(S, Result div 2);
        JetError(JetRetrieveColumn(Table.Database.Session.Handle, Table.Handle, Handle, @S[1], Result, Result, 0, nil));
        string(Value) := S;
      end
      else
        string(Value) := '';
    end;
  else
    begin
      JetError(JetRetrieveColumn(Table.Database.Session.Handle, Table.Handle, Handle, nil, 0, Result, 0, nil));
      JetError(JetRetrieveColumn(Table.Database.Session.Handle, Table.Handle, Handle, @Value, Result, Result, 0, nil));
    end;
  end;
end;

procedure TMSJetColumn.Update(const P: Pointer; Size: Cardinal);
begin
  JetError(JetSetColumn(Table.Database.Session.Handle, Table.Handle, Handle, P, Size, 0, nil));
end;

function TMSJetColumn.GetIsNull: Boolean;
var
  Size: ULONG;
  Buf: array[0..255] of AnsiChar;
begin
  Size := 0;
  JetError(JetRetrieveColumn(Table.Database.Session.Handle, Table.Handle, Handle, @Buf[0], 256, Size, 0, nil));
  Result := Size = 0;
end;

procedure TMSJetColumn.SetHandle(const Value: JET_COLUMNID);
begin
  if FHandle <> Value then
  begin
    FHandle := Value;

    DbReload;
  end;
end;

function TMSJetColumn.GetAsBoolean: Boolean;
begin
  Result := AsInt64 > 0;
end;

function TMSJetColumn.GetAsCurrency: Currency;
begin
  Result := AsInt64;
end;

function TMSJetColumn.GetAsDateTime: TDateTime;
begin
  if IsNull then
  begin
    Result := 0;
    Exit;
  end;

  if DataType = jetDateTime then
    Retrieve(Result)
  else
    Result := AsFloat;
end;

function TMSJetColumn.GetAsFloat: Double;
begin
  if IsNull then
  begin
    Result := 0;
    Exit;
  end;

  case DataType of
  jetDouble:   Retrieve(Result);
  jetSingle:   Result := AsSingle;
  jetCurrency: Result := AsCurrency;
  else
    Result := AsInt64;
  end;
end;

function TMSJetColumn.GetAsInteger: Integer;
var
  U1: Byte;
  U2: Word;
  S2: SmallInt;
  U4: Cardinal;
  S4: Longint;
begin
  if IsNull then
  begin
    Result := 0;
    Exit;
  end;

  case DataType of
    jetBit, jetUByte:
      begin
        Retrieve(U1);
        Result := U1;
      end;
    jetShort:
      begin
        Retrieve(S2);
        Result := S2;
      end;
    jetUShort:
      begin
        Retrieve(U2);
        Result := U2;
      end;
    jetLong:
      begin
        Retrieve(S4);
        Result := S4;
      end;
    jetULong:
      begin
        Retrieve(U4);
        Result := U4;
      end;
  else
    raise EMSJetColumnTypeError.Create(Name, DataType);
  end;
end;

function TMSJetColumn.GetAsRaw: RawByteString;
var
  Size: ULONG;
begin
  if IsNull then
  begin
    Result := '';
    Exit;
  end;

  case DataType of
    jetText,
    jetLongText: Retrieve(Result);
    jetBit,
    jetUByte,
    jetShort,
    jetLong,
    jetULong,
    jetUShort:   Result := RawByteString(IntToStr(AsInteger));
    jetLongLong: Result := RawByteString(IntToStr(AsInt64));
    jetCurrency: Result := RawByteString(CurrToStr(AsCurrency));
    jetSingle,
    jetDouble:   Result := RawByteString(FloatToStr(AsFloat));
    jetDateTime: Result := RawByteString(DateTimeToStr(AsDateTime));
  else
    JetError(JetRetrieveColumn(Table.Database.Session.Handle, Table.Handle, Handle, nil, 0, Size, 0, nil));
    if Size > 0 then
    begin
      SetLength(Result, Size);
      JetError(JetRetrieveColumn(Table.Database.Session.Handle, Table.Handle, Handle, @Result[1], Size, Size, 0, nil));
    end;
  end;
end;

function TMSJetColumn.GetAsInt64: Int64;
begin
  if IsNull then
  begin
    Result := 0;
    Exit;
  end;

  if DataType = jetLongLong then
    Retrieve(Result)
  else
    Result := AsInteger;
end;

function TMSJetColumn.GetAsSingle: Single;
begin
  if IsNull then
  begin
    Result := 0;
    Exit;
  end;

  if DataType = jetSingle then
    Retrieve(Result)
  else
    Result := AsFloat;
end;

function TMSJetColumn.GetAsString: string;
var
  Size: ULONG;
begin
  if IsNull then
  begin
    Result := '';
    Exit;
  end;

  case DataType of
    jetText,
    jetLongText: Retrieve(Result);
    jetBit:      Result := BoolToStr(AsBoolean, True);
    jetUByte,
    jetShort,
    jetLong,
    jetULong,
    jetUShort:   Result := IntToStr(AsInteger);
    jetLongLong: Result := IntToStr(AsInt64);
    jetCurrency: Result := CurrToStr(AsCurrency);
    jetSingle,
    jetDouble:   Result := FloatToStr(AsFloat);
    jetDateTime: Result := DateTimeToStr(AsDateTime);
  else
    JetError(JetRetrieveColumn(Table.Database.Session.Handle, Table.Handle, Handle, nil, 0, Size, 0, nil));
    if Size > 0 then
    begin
{      SetLength(S, Size);
      JetError(JetRetrieveColumn(Table.Database.Session.Handle, Table.Handle, Handle, @S[1], Size, Size, 0, nil));
      Result := string(S);}
      SetLength(Result, Size div 2);
      JetError(JetRetrieveColumn(Table.Database.Session.Handle, Table.Handle, Handle, @Result[1], Size, Size, 0, nil));
    end;
  end;
end;

function TMSJetColumn.GetValue: Variant;
begin
  if IsNull then
  begin
    Result := Null;
    Exit;
  end;

  case DataType of
    jetText,
    jetLongText: Result := AsString;
    jetBit:      Result := AsBoolean;
    jetUByte,
    jetShort,
    jetLong,
    jetULong,
    jetUShort:   Result := AsInteger;
    jetLongLong: Result := AsInt64;
    jetCurrency: Result := AsCurrency;
    jetSingle,
    jetDouble:   Result := AsFloat;
    jetDateTime: Result := AsDateTime;
  else
    raise EMSJetColumnTypeError.Create(Name, DataType);
  end;
end;

procedure TMSJetColumn.SetAsString(const Value: string);
begin
  if Length(Value) = 0 then
  begin
    IsNull := True;
    Exit;
  end;

  case DataType of
    jetBit:      AsBoolean := Value = '1';
    jetUByte,
    jetShort,
    jetLong,
    jetULong,
    jetUShort:   AsInteger := StrToInt(string(Value));
    jetLongLong: AsInt64 := StrToInt64(string(Value));
    jetCurrency: AsCurrency := StrToCurr(string(Value));
    jetSingle,
    jetDouble:   AsFloat := StrToFloat(string(Value));
    jetDateTime: AsDateTime := StrToDateTime(string(Value));
  else
    begin
{      S := AnsiString(Value);
      Update(@S[1], Length(S));}
      Update(@Value[1], Length(Value) * 2);
    end;
  end;
end;

procedure TMSJetColumn.SetDisplayName(const Value: string);
begin
  FDisplayName := Value;
end;

procedure TMSJetColumn.SetAsInteger(const Value: Integer);

  procedure U1(const Value: Byte);
  begin
    Update(@Value, 1);
  end;

  procedure U2(const Value: Word);
  begin
    Update(@Value, 2);
  end;

  procedure S2(const Value: SmallInt);
  begin
    Update(@Value, 2);
  end;

  procedure U4(const Value: Cardinal);
  begin
    Update(@Value, 4);
  end;

  procedure S4(const Value: Integer);
  begin
    Update(@Value, 4);
  end;

  procedure S8(const Value: Int64);
  begin
    Update(@Value, 8);
  end;

begin
  case DataType of
    jetBit, jetUByte:
      U1(Value);
    jetShort:
      S2(Value);
    jetUShort:
      U2(Value);
    jetLong:
      S4(Value);
    jetULong:
      U4(Value);
    jetLongLong:
      S8(Value);
  else
    raise EMSJetColumnTypeError.Create(Name, DataType);
  end;
end;

procedure TMSJetColumn.SetAsRaw(const Value: RawByteString);
begin
if Length(Value) = 0 then
  begin
    IsNull := True;
    Exit;
  end;

  case DataType of
    jetBit:      AsBoolean := Value = '1';
    jetUByte,
    jetShort,
    jetLong,
    jetULong,
    jetUShort:   AsInteger := StrToInt(string(Value));
    jetLongLong: AsInt64 := StrToInt64(string(Value));
    jetCurrency: AsCurrency := StrToCurr(string(Value));
    jetSingle,
    jetDouble:   AsFloat := StrToFloat(string(Value));
    jetDateTime: AsDateTime := StrToDateTime(string(Value));
  else
    Update(@Value[1], Length(Value));
  end;
end;

procedure TMSJetColumn.SetAsBoolean(const Value: Boolean);
begin
  if Value then
    AsInteger := 1
  else
    AsInteger := 0;
end;

procedure TMSJetColumn.SetAsCurrency(const Value: Currency);
begin
  if DataType = jetCurrency then
    Update(@Value, SizeOf(Value))
  else
    AsFloat := Value;
end;

procedure TMSJetColumn.SetAsDateTime(const Value: TDateTime);
begin
  if DataType = jetDateTime then
    Update(@Value, SizeOf(Value))
  else
    AsFloat := Value;
end;

procedure TMSJetColumn.SetAsFloat(const Value: Double);
begin
  case DataType of
    jetDouble:   Update(@Value, SizeOf(Value));
    jetSingle:   AsSingle := Value;
    jetCurrency: AsCurrency := Value;
  else
    raise EMSJetColumnTypeError.Create(Name, DataType);
  end;
end;

procedure TMSJetColumn.SetAsInt64(const Value: Int64);
begin
  if DataType = jetLongLong then
    Update(@Value, SizeOf(Value))
  else
    AsFloat := Value;
end;

procedure TMSJetColumn.SetAsSingle(const Value: Single);
begin
  if DataType = jetSingle then
    Update(@Value, SizeOf(Value))
  else
    AsFloat := Value;
end;

procedure TMSJetColumn.SetIsNull(const Value: Boolean);
begin
  JetError(JetSetColumn(Table.Database.Session.Handle, Table.Handle, Handle, nil, 0, 0, nil));
end;

procedure TMSJetColumn.SetValue(const Value: Variant);
begin
  if VarIsNull(Value) or VarIsClear(Value) then
  begin
    IsNull := True;
    Exit;
  end;

  case DataType of
    jetText,
    jetLongText: AsString := VarToStr(Value);
    jetBit:      AsBoolean := Value;
    jetUByte,
    jetShort,
    jetLong,
    jetULong,
    jetUShort:   AsInteger := Value;
    jetLongLong: AsInt64 := Value;
    jetCurrency: AsCurrency := Value;
    jetSingle,
    jetDouble:   AsFloat := Value;
    jetDateTime: AsDateTime := Value;
  else
    raise EMSJetColumnTypeError.Create(Name, DataType);
  end;
end;

{ TMSJetColumnList }

function TMSJetColumnList.Add(const Name: string; DataType: TMSJetColumnType; Size: Cardinal; Options: TMSJetColumnOptions): TMSJetColumn;
begin
  Result := Add(Name);
  Result.DataType := DataType;
  Result.Size := Size;
  Result.Options := Options;
end;

function  TMSJetColumnList.CreateItem(const Name: string): TMSJetColumn;
begin
  Result := TMSJetColumn.Create(Table);
  Result.FName := AnsiString(Name);
end;

procedure TMSJetColumnList.Reload;
var
  List: JET_COLUMNLIST;
  Temp: TMSJetTable;
  Rest: TList;
  I: Integer;
  BufSize: ULONG;
  Column: TMSJetColumn;
begin
  Rest := TList.Create;
  try
    for I := 0 to Count - 1 do
      Rest.Add(Items[I]);

    ZeroMemory(@List, SizeOf(List));
    List.cbStruct := SizeOf(List);
    JetError(JetGetTableColumnInfo(Table.Database.Session.Handle, Table.Handle, '', @List, SizeOf(JET_COLUMNLIST), 1));
    Temp := TMSJetTable.Create(Table.Database, List.tableid);
    try
      if Temp.Last then
      repeat
        Column := FindOrAdd(Temp.GetColumnText(List.columnname));
        Column.ItemState := itsUnmodified;
        Rest.Remove(Column);

        JetError(JetRetrieveColumn(Table.Database.Session.Handle, Temp.Handle, List.columnid, @Column.FHandle, SizeOf(Column.FHandle), BufSize, 0, nil));
        JetError(JetRetrieveColumn(Table.Database.Session.Handle, Temp.Handle, List.coltyp, @Column.FDataType, SizeOf(JET_COLTYP), BufSize, 0, nil));
        JetError(JetRetrieveColumn(Table.Database.Session.Handle, Temp.Handle, List.cbMax, @Column.FSize, SizeOf(Column.FSize), BufSize, 0, nil));
        JetError(JetRetrieveColumn(Table.Database.Session.Handle, Temp.Handle, List.grbit, @Column.FOptions, SizeOf(Column.FHandle), BufSize, 0, nil));
      until not Temp.Previous;
    finally
      Temp.Free;
    end;

    for I := 0 to Rest.Count - 1 do
      TMSJetColumn(Rest[I]).Free;
  finally
    Rest.Free;
  end;
end;

{ TMSJetIndex }

constructor TMSJetIndex.Create(Table: TMSJetTable);
begin
  inherited Create(Table);
  FDensity := 80;
  FColumns := TList.Create;
end;

destructor TMSJetIndex.Destroy;
begin
  FreeAndNil(FColumns);
  inherited;
end;

procedure TMSJetIndex.DoDbDelete(const Name: AnsiString);
begin
  JetError(JetDeleteIndex(Table.Database.Session.Handle, Table.Handle, PAnsiChar(Name)));
  FHandle := 0;
end;

procedure TMSJetIndex.DoDbRename(const OldName, NewName: AnsiString);
begin
  DoDbDelete(OldName);
  DoDbCreate(NewName);
end;

procedure TMSJetIndex.DoDbUpdate(const Name: AnsiString);
begin
  DoDbDelete(Name);
  DoDbCreate(Name);
end;

function TMSJetIndex.GetColumn(const Index: Integer): TMSJetColumn;
begin
  Result := TMSJetColumn(FColumns[Index]);
end;

function TMSJetIndex.GetColumnCount: Integer;
begin
  Result := FColumns.Count;
end;

procedure TMSJetIndex.DoDbReload(const Name: AnsiString);
var
  List: JET_INDEXLIST;
  Temp: TMSJetTable;
  Column: string;
  L: ULONG;
  BufSize: ULONG;
  S: string;
begin
  ZeroMemory(@List, SizeOf(List));
  List.cbStruct := SizeOf(List);
  JetError(JetGetTableIndexInfo(Table.Database.Session.Handle, Table.Handle, PAnsiChar(Name), @List, SizeOf(List), 0));

  Temp := TMSJetTable.Create(Table.Database, List.tableid);
  try
    S := '';
    if Temp.First then
    repeat
      ItemState := itsUnmodified;
      Handle := 1; // Fake handle, because ESE have not handles for indexes

      JetError(JetRetrieveColumn(Table.Database.Session.Handle, Temp.Handle, List.columnidgrbitIndex, @FOptions, 4, BufSize, 0, nil));

      Column := Temp.GetColumnText(List.columnidcolumnname);
      JetError(JetRetrieveColumn(Table.Database.Session.Handle, Temp.Handle, List.columnidgrbitColumn, @L, SizeOf(L), BufSize, 0, nil));

      if S <> '' then
        S := S + ',';
      if L = 1 then
        S := S + '-'
      else
        S := S + '+';
      S := S + Column;
    until not Temp.Next;
    Fields := S;
  finally
    Temp.Free;
  end;
end;

procedure TMSJetIndex.DoDbCreate(const Name: AnsiString);
begin
  JetError(JetCreateIndex(Table.Database.Session.Handle, Table.Handle, PAnsiChar(FName), SetToInt(Options), PAnsiChar(FKey), Length(FKey), Density));
  FHandle := 0;
end;

procedure TMSJetIndex.AssignTo(Dest: TPersistent);
begin
  if Dest is TMSJetIndex then
  begin
    TMSJetIndex(Dest).FName := FName;
    TMSJetIndex(Dest).FFields := Fields;
    TMSJetIndex(Dest).FDensity := Density;
    TMSJetIndex(Dest).FOptions := Options;
  end
  else
    inherited;
end;

procedure TMSJetIndex.SetFields(const Value: string);
var
  I: Integer;
  List: TStringList;
begin
  if FFields <> Value then
  begin
    FFields := Value;
    FKey := AnsiString(Value);
    for I := 1 to Length(FKey) do
      if FKey[I] = ',' then
        FKey[I] := #0;
    FKey := FKey + #0#0;

    FColumns.Clear;
    List := TStringList.Create;
    List.CommaText := Value;
    for I := 0 to List.Count - 1 do
    begin
      FColumns.Add(Table.Columns.ByName(Copy(List[I], 2, MaxInt)));
    end;
  end;
end;

{ TMSJetIndexList }

function TMSJetIndexList.Add(const Name, Columns: string; Options: TMSJetIndexOptions): TMSJetIndex;
begin
  Result := Add(Name);
  Result.Fields := Columns;
  Result.Options := Options;
end;

function  TMSJetIndexList.CreateItem(const Name: string): TMSJetIndex;
begin
  Result := TMSJetIndex.Create(Table);
  Result.FName := AnsiString(Name);
end;

procedure TMSJetIndexList.Reload;
var
  List: JET_INDEXLIST;
  Temp: TMSJetTable;
  Index: TMSJetIndex;
  S: string;
begin
  Clear;

  ZeroMemory(@List, SizeOf(List));
  List.cbStruct := SizeOf(List);
  JetError(JetGetTableIndexInfo(Table.Database.Session.Handle, Table.Handle, nil, @List, SizeOf(List), 1));

  Temp := TMSJetTable.Create(Table.Database, List.tableid);
  try
    S := '';
    if Temp.Last then
    repeat
      Index := TMSJetIndex(FindOrAdd(Temp.GetColumnText(List.columnidindexname)));
      Index.DbReload;
    until not Temp.Previous;
  finally
    Temp.Free;
  end;
end;

{ TMSJetJoin }

constructor TMSJetJoin.Create(Parent: TMSJetTable; const TableName, IndexName, ParentKey: string);
begin
  inherited Create;
  FParent := Parent;
  FParentKey := ParentKey;

  FTable := TMSJetTable.Create(Parent.Database, TableName);
  FTable.IndexName := IndexName;
  FTable.Open;

  Parent.FJoins.Add(ParentKey, Self);
end;

destructor TMSJetJoin.Destroy;
begin
  if Parent <> nil then
    Parent.FJoins.Remove(ParentKey);

  FreeAndNil(FTable);
  inherited;
end;

function TMSJetJoin.GetValue(const Name: string): Variant;
begin
  if IsSync then
    Result := Table[Name].Value
  else
    Result := Null;
end;

function TMSJetJoin.Sync(IsSync: Boolean): Boolean;
begin
  if IsSync then
    FIsSync := Table.Seek([Parent[ParentKey].Value], jetSeekEQ, True)
  else
    FIsSync := False;

  Result := FIsSync;

  Table.SyncCursor(Result);
end;

{ TMSJetTable }

constructor TMSJetTable.Create(Database: TMSJetDatabase; const Name: string);
begin
  inherited Create;

  FName := AnsiString(Name);
  FDatabase := Database;
  FColumns := TMSJetColumnList.Create(Self);
  FIndexes := TMSJetIndexList.Create(Self);
  FJoins := TDictionary<string, TMSJetJoin>.Create;
  FPages := 2;
  FDensity := 80;
end;

constructor TMSJetTable.Create(Database: TMSJetDatabase; Handle: JET_TABLEID);
begin
  inherited Create;

  FDatabase := Database;
  FHandle := Handle;
  FColumns := TMSJetColumnList.Create(Self);
  FIndexes := TMSJetIndexList.Create(Self);
  FJoins := TDictionary<string, TMSJetJoin>.Create;
end;

destructor TMSJetTable.Destroy;
var
  Item: TPair<string, TMSJetJoin>;
begin
  Close;
  if FKey <> nil then
    FreeMem(FKey);
  FreeAndNil(FColumns);
  FreeAndNil(FIndexes);

  for Item in FJoins do
    Item.Value.Free;

  FreeAndNil(FJoins);

  inherited;
end;

procedure TMSJetTable.Open;
begin
  Close;

  JetError(JetOpenTable(Database.Session.Handle, Database.Handle, PAnsiChar(FName), nil, 0, JET_bitTableUpdatable, FHandle));

  Columns.Reload;
  Indexes.Reload;

  if IndexName <> '' then
    JetError(JetSetCurrentIndex(Database.Session.Handle, Handle, PAnsiChar(FIndexName)));

  First;
end;

procedure TMSJetTable.Close;
begin
  if FHandle <> 0 then
  begin
    JetError(JetCloseTable(Database.Session.Handle, FHandle));
    FHandle := 0;
  end;
end;

function TMSJetTable.First: Boolean;
var
  Err: JET_ERR;
begin
  Err := JetMove(Database.Session.Handle, Handle, JET_MoveFirst, 0);
  if Err = 0 then
    Result := True
  else if Err = JET_errNoCurrentRecord then
    Result := False
  else
  begin
    JetError(Err);
    Result := False;
  end;

  SyncCursor(Result);
end;

function TMSJetTable.Last: Boolean;
var
  Err: JET_ERR;
begin
  Err := JetMove(Database.Session.Handle, Handle, JET_MoveLast, 0);
  if Err = 0 then
    Result := True
  else if Err = JET_errNoCurrentRecord then
    Result := False
  else
  begin
    JetError(Err);
    Result := False;
  end;

  SyncCursor(Result);
end;

function TMSJetTable.Next: Boolean;
var
  Err: JET_ERR;
begin
  Err := JetMove(Database.Session.Handle, Handle, JET_MoveNext, 0);
  if Err = 0 then
    Result := True
  else if Err = JET_errNoCurrentRecord then
    Result := False
  else
  begin
    JetError(Err);
    Result := False;
  end;

  SyncCursor(Result);
end;

function TMSJetTable.Previous: Boolean;
var
  Err: JET_ERR;
begin
  Err := JetMove(Database.Session.Handle, Handle, JET_MovePrevious, 0);
  if Err = 0 then
    Result := True
  else if Err = JET_errNoCurrentRecord then
    Result := False
  else
  begin
    JetError(Err);
    Result := False;
  end;

  SyncCursor(Result);
end;

procedure TMSJetTable.Delete;
begin
  JetError(JetDelete(Database.Session.Handle, Handle));
end;

procedure TMSJetTable.SetColumns(const Value: TMSJetColumnList);
begin
  FColumns.Assign(Value);
end;

procedure TMSJetTable.SetIndexes(const Value: TMSJetIndexList);
begin
  FIndexes.Assign(Value);
end;

procedure TMSJetTable.Cancel;
begin
  JetError(JetPrepareUpdate(Database.Session.Handle, Handle, JET_prepCancel));
  FIsInsert := False;
end;

procedure TMSJetTable.Insert;
begin
  JetError(JetPrepareUpdate(Database.Session.Handle, Handle, JET_prepInsert));
  FIsInsert := True;
end;

function TMSJetTable.Join(const TableName, IndexName, ParentKey: string): TMSJetJoin;
begin
  Result := TMSJetJoin.Create(Self, TableName, IndexName, ParentKey);
end;

procedure TMSJetTable.Edit;
begin
  JetError(JetPrepareUpdate(Database.Session.Handle, Handle, JET_prepReplace));
  FIsInsert := False;
end;

procedure TMSJetTable.Post;
var
  Size: ULONG;
  Bookmark: array[0..1024] of Byte;
begin
  JetError(JetUpdate(Database.Session.Handle, Handle, @Bookmark[0], 1024, Size));
  if IsInsert then
  begin
    JetError(JetGotoBookmark(Database.Session.Handle, Handle, @Bookmark[0], Size));
    FIsInsert := False;
  end;
end;

procedure TMSJetTable.SetIndexName(const Value: string);
var
  S: AnsiString;
begin
  FIndexName := AnsiString(Value);
  if Handle <> 0 then
  begin
    S := AnsiString(Value);
    JetError(JetSetCurrentIndex(Database.Session.Handle, Handle, PAnsiChar(S)));
    First;
  end;
end;

procedure TMSJetTable.SyncCursor(HasCursor: Boolean);
var
  Item: TPair<string, TMSJetJoin>;
begin
  for Item in FJoins do
    Item.Value.Sync(HasCursor);
end;

procedure TMSJetTable.ClearRange;
begin
  JetError(JetSetIndexRange(Database.Session.Handle, Handle, JET_bitRangeRemove));
  FSeekSetRange := False;
end;

function TMSJetTable.Seek(const Keys: array of Variant; Option: TMSJetSeekOption = jetSeekEQ; SetRange: Boolean = False): Boolean;
var
  I: Integer;
  Index: TMSJetIndex;
  grbit: ULONG;
  Err: JET_ERR;
begin
  Index := nil;
  if IndexName = '' then
  begin
    for I := 0 to Indexes.Count - 1 do
      if jetIndexPrimary in Indexes[I].Options then
      begin
        Index := Indexes[I];
        Break;
      end;
  end
  else
    Index := Indexes.ByName(IndexName);

  MakeKey(Index.Columns[0], Keys[0], [jetNewKey]);
  for I := 1 to High(Keys) do
    MakeKey(Index.Columns[I], Keys[I], []);

  for I := High(Keys) + 1 to Index.ColumnCount - 1 do
    MakeKey(Index.Columns[I], Null, []);

  grbit := Cardinal(Option);
  if SetRange then
    grbit := grbit or JET_bitSetIndexRange;

  Err := JetSeek(Database.Session.Handle, Handle,  grbit);
  Result := Err >= 0;
  if not Result and (Err <> JET_errRecordNotFound) then
    JetError(Err);

  SetLength(FSeekKeys, High(Keys) + 1);
  for I := 0 to High(Keys) do
    FSeekKeys[I] := Keys[I];

  FSeekOption := Option;
  FSeekSetRange := SetRange;

  SyncCursor(Result);
end;

function TMSJetTable.RetrieveKey: Boolean;
var
  Size: ULONG;
  Err: JET_ERR;
begin
  Err := JetRetrieveKey(Database.Session.Handle, Handle, nil, 0, Size, 0);
  Result := Err = 0;
  if Err <> JET_errNoCurrentRecord then
  begin
    JetError(Err);
    ReallocMem(FKey, Size);
    JetRetrieveKey(Database.Session.Handle, Handle, FKey, Size, Size, 0);
    FKeySize := Size;
  end;
end;

function TMSJetTable.CreateBookmark: JetBookmark;
begin
  if IndexName = '' then
    Result := CreatePrimaryBookmark
  else
    Result := CreateSecondaryBookmark;
end;

function TMSJetTable.CreatePrimaryBookmark: JetBookmark;
var
  Data: Pointer;
  Size: ULONG;
  Err: JET_ERR;
begin
  Err := JetGetBookmark(Database.Session.Handle, Handle, nil, 0, Size);
  case Err of
    JET_errNoCurrentRecord:
    begin
      if First then
      begin
        Err := JetGetBookmark(Database.Session.Handle, Handle, nil, 0, Size);
        if Err <> JET_errBufferTooSmall then
          JetError(Err);
      end
      else
      begin
        Result := nil;
        Exit;
      end;
    end;
    JET_errBufferTooSmall:
      ;
    else
      JetError(Err);
  end;

  GetMem(Data, Size);
  JetError(JetGetBookmark(Database.Session.Handle, Handle, Data, Size, Size));
  Result := TJetBookmark.Create(Self, Data, Size, nil, 0);
end;

function TMSJetTable.CreateSecondaryBookmark: JetBookmark;
var
  Data1, Data2: Pointer;
  Size1, Size2: ULONG;
  Err: JET_ERR;
begin
  Err := JetGetSecondaryIndexBookmark(Database.Session.Handle, Handle, nil, 0, Size1, nil, 0, Size2, 0);
  case Err of
    JET_errNoCurrentRecord:
    begin
      if First then
      begin
        Err := JetGetSecondaryIndexBookmark(Database.Session.Handle, Handle, nil, 0, Size1, nil, 0, Size2, 0);
        if Err <> JET_errBufferTooSmall then
          JetError(Err);
      end
      else
      begin
        Result := nil;
        Exit;
      end;
    end;
    JET_errBufferTooSmall:
      ;
    else
      JetError(Err);
  end;

  GetMem(Data1, Size1);
  GetMem(Data2, Size2);
  JetError(JetGetSecondaryIndexBookmark(Database.Session.Handle, Handle, Data1, Size1, Size1, Data2, Size2, Size2, 0));
  Result := TJetBookmark.Create(Self, Data1, Size1, Data2, Size2);
end;

function  TMSJetTable.GotoBookmark(const Bookmark: JetBookmark): Boolean;
begin
  Result := Bookmark.Go;
  SyncCursor(Result);
end;

function  TMSJetTable.GoToKey: Boolean;
var
  Err: JET_ERR;
begin
  JetError(JetMakeKey(Database.Session.Handle, Handle, FKey, FKeySize, JET_bitNormalizedKey));
  Err := JetSeek(Database.Session.Handle, Handle, JET_GRBIT(jetSeekEQ));
  if Err = 0 then
    Result := True
  else if Err = JET_errNoCurrentRecord then
    Result := False
  else
  begin
    JetError(Err);
    Result := False;
  end;
  SyncCursor(Result);
end;

procedure TMSJetTable.MakeKey(Column: TMSJetColumn; const Value: Variant; Options: TMSJetKeyOptions);

  procedure MakeBool(const Value: Boolean);
  begin
    JetError(JetMakeKey(Database.Session.Handle, Handle, @Value, SizeOf(Value), SetToInt(Options)));
  end;

  procedure MakeByte(const Value: Byte);
  begin
    JetError(JetMakeKey(Database.Session.Handle, Handle, @Value, SizeOf(Value), SetToInt(Options)));
  end;

  procedure MakeSmallint(const Value: SmallInt);
  begin
    JetError(JetMakeKey(Database.Session.Handle, Handle, @Value, SizeOf(Value), SetToInt(Options)));
  end;

  procedure MakeInteger(const Value: Integer);
  begin
    JetError(JetMakeKey(Database.Session.Handle, Handle, @Value, SizeOf(Value), SetToInt(Options)));
  end;

  procedure MakeCurrency(const Value: Currency);
  begin
    JetError(JetMakeKey(Database.Session.Handle, Handle, @Value, SizeOf(Value), SetToInt(Options)));
  end;

  procedure MakeSingle(const Value: Single);
  begin
    JetError(JetMakeKey(Database.Session.Handle, Handle, @Value, SizeOf(Value), SetToInt(Options)));
  end;

  procedure MakeDouble(const Value: Double);
  begin
    JetError(JetMakeKey(Database.Session.Handle, Handle, @Value, SizeOf(Value), SetToInt(Options)));
  end;

  procedure MakeDateTime(const Value: TDateTime);
  begin
    JetError(JetMakeKey(Database.Session.Handle, Handle, @Value, SizeOf(Value), SetToInt(Options)));
  end;

  procedure MakeCardinal(const Value: Cardinal);
  begin
    JetError(JetMakeKey(Database.Session.Handle, Handle, @Value, SizeOf(Value), SetToInt(Options)));
  end;

  procedure MakeInt64(const Value: Int64);
  begin
    JetError(JetMakeKey(Database.Session.Handle, Handle, @Value, SizeOf(Value), SetToInt(Options)));
  end;

  procedure MakeWord(const Value: Word);
  begin
    JetError(JetMakeKey(Database.Session.Handle, Handle, @Value, SizeOf(Value), SetToInt(Options)));
  end;

  procedure MakeBinary(const Value: RawByteString);
  begin
    JetError(JetMakeKey(Database.Session.Handle, Handle, PAnsiChar(Value), Length(Value), SetToInt(Options)));
  end;

  procedure MakeText(const Value: string);
  begin
    JetError(JetMakeKey(Database.Session.Handle, Handle, PWideChar(Value), Length(Value) * 2, SetToInt(Options)));
  end;

begin
  case Column.DataType of
    jetBit:
      MakeBool(Value);
    jetUByte:
      MakeByte(Value);
    jetShort:
      MakeSmallInt(Value);
    jetLong:
      MakeInteger(Value);
    jetCurrency:
      MakeCurrency(Value);
    jetSingle:
      MakeSingle(Value);
    jetDouble:
      MakeDouble(Value);
    jetDateTime:
      MakeDateTime(Value);
    jetBinary, jetLongBinary:
      MakeBinary(RawByteString(Value));
    jetText, jetLongText:
      MakeText(VarToStr(Value));
    jetULong:
      MakeCardinal(Value);
    jetLongLong:
      MakeInt64(Value);
    jetUShort:
      MakeWord(Value);
  end;
end;

function TMSJetTable.AddColumn(const Name: string; DataType: TMSJetColumnType; Size: ULONG; Options: TMSJetColumnOptions): TMSJetColumn;
var
  Def: JET_COLUMNDEF;
begin
  Result := Columns.Add(Name);
  try
    Result.FDataType := DataType;
    Result.FSize := Size;
    Result.FOptions := Options;

    ZeroMemory(@Def, SizeOf(Def));
    Def.cbStruct := SizeOf(Def);
    Def.coltyp := ULONG(DataType);
    Def.cp := 1200;
    Def.cbMax := Size;
    Def.grbit := SetToInt(Options);

    JetError(JetAddColumn(Database.Session.Handle, Handle, PAnsiChar(Result.FName), @Def, nil, 0, Result.FHandle));
  except
    FreeAndNil(Result);
    raise;
  end;
end;

function TMSJetTable.GetColumn(const ColumnName: string): TMSJetColumn;
begin
  Result := Columns.ByName(ColumnName);
end;

function TMSJetTable.GetColumnText(id: JET_COLUMNID): string;
var
  Buf: array[0..255] of AnsiChar;
  BufSize: ULONG;
begin
  JetError(JetRetrieveColumn(Database.Session.Handle, Handle, id, @Buf[0], 255, BufSize, 0, nil));
  Result := string(AnsiString(Copy(Buf, 1, BufSize)));
end;

function TMSJetTable.GetIndexName;
begin
  Result := string(FIndexName);
end;

function TMSJetTable.GetJoins: TArray<TMSJetJoin>;
begin
  Result := FJoins.Values.ToArray;
end;

function TMSJetTable.GetJoinCount: Integer;
begin
  Result := FJoins.Count;
end;

function TMSJetTable.GetLink(const ParentKey: string): TMSJetJoin;
begin
  Result := FJoins[ParentKey];
end;

procedure TMSJetTable.DoDbDelete(const Name: AnsiString);
begin
  Close;
  JetError(JetDeleteTable(Database.Session.Handle, Database.Handle, PAnsiChar(AnsiString(Name))));
end;

procedure TMSJetTable.DoDbCreate(const Name: AnsiString);
type
  JET_COLUMNCREATE_ARRAY = array[0..MaxInt div SizeOf(JET_COLUMNCREATE) - 1] of JET_COLUMNCREATE;
  JET_INDEXCREATE_ARRAY = array[0..MaxInt div SizeOf(JET_INDEXCREATE) - 1] of JET_INDEXCREATE;
var
  TableCreate: JET_TABLECREATE;
  Cols: ^JET_COLUMNCREATE_ARRAY;
  Inds: ^JET_INDEXCREATE_ARRAY;
  I: Integer;
  Unicode: JET_UNICODEINDEX;
begin
  Unicode.lcid := GetSystemDefaultLCID;
  Unicode.dwMapFlags := 0;//LCMAP_SORTKEY or NORM_IGNORECASE or NORM_IGNOREKANATYPE or NORM_IGNOREWIDTH;

  Cols := AllocMem(SizeOf(JET_COLUMNCREATE) * Columns.Count);
  Inds := AllocMem(SizeOf(JET_INDEXCREATE) * Indexes.Count);
  try
    for I := 0 to Columns.Count - 1 do
    begin
      Cols[I].cbStruct := SizeOf(JET_COLUMNCREATE);
      Cols[I].szColumnName := PAnsiChar(FColumns[I].FName);
      Cols[I].coltyp := Cardinal(Columns[I].DataType);
      Cols[I].cbMax := Columns[I].Size;
      Cols[I].grbit := SetToInt(Columns[I].Options);
      //Cols[I].pvDefault := Columns[I];
      //Cols[I].cbDefault := Columns[I];
      Cols[I].cp := 1200; // Unicode
    end;

    for I := 0 to Indexes.Count - 1 do
    begin
      Inds[I].cbStruct := SizeOf(JET_INDEXCREATE);
      Inds[I].szIndexName := PAnsiChar(Indexes[I].FName);
      Inds[I].szKey := PAnsiChar(Indexes[I].FKey);
      Inds[I].cbKey := Length(Indexes[I].FKey);
      Inds[I].ulDensity := Indexes[I].Density;
      Inds[I].grbit := SetToInt(Indexes[I].Options);
      Inds[I].lcid := GetSystemDefaultLCID;
{      if jetIndexUnicode in Indexes[I].Options then
        Pointer(Inds[I].UNICODE) := @Unicode
      else
        ULONG(Inds[I].UNICODE) := GetSystemDefaultLCID;}
    end;

    ZeroMemory(@TableCreate, SizeOf(TableCreate));
    TableCreate.cbStruct := SizeOf(TableCreate);
    TableCreate.szTableName := PAnsiChar(Name);
    TableCreate.ulPages := Pages;
    TableCreate.ulDensity := Density;
    TableCreate.cColumns := Columns.Count;
    TableCreate.rgcolumncreate := Pointer(Cols);
    TableCreate.cIndexes := Indexes.Count;
    TableCreate.rgindexcreate := Pointer(Inds);

    Database.Session.Instance.BeginCall;
    try
      JetError(JetCreateTableColumnIndex(Database.Session.Handle, Database.Handle, TableCreate), Database.Session.Instance.Handle, Database.Session.Handle);
    finally
      Database.Session.Instance.EndCall;
      FHandle := TableCreate.tableid;
      Columns.Reload;
      Indexes.Reload;
    end;
  finally
    FreeMem(Cols);
    FreeMem(Inds);
  end;
end;

procedure TMSJetTable.DoDbUpdate(const Name: AnsiString);
begin
  Columns.Apply;
  Indexes.Apply;
end;

procedure TMSJetTable.DoDbRename(const OldName, NewName: AnsiString);
begin
  JetError(JetRenameTable(Database.Session.Handle, Database.Handle, PAnsiChar(OldName), PAnsiChar(NewName)));
end;

procedure TMSJetTable.DoDbReload(const Name: AnsiString);
begin
  Columns.Reload;
  Indexes.Reload;
end;


{ TMSJetTableList }

constructor TMSJetTableList.Create(Database: TMSJetDatabase);
begin
  inherited Create;
  FDatabase := Database;
end;

function TMSJetTableList.CreateItem(const Name: string): TMSJetTable;
begin
  Result := TMSJetTable.Create(Database, Name);
end;

{ TJetBookmark }

constructor TJetBookmark.Create(Table: TMSJetTable; const Data: Pointer; const Size: ULONG; const SecondaryData: Pointer; const SecondarySize: ULONG);
begin
  FTable := Table;
  FData := Data;
  FSize := Size;
  FSecondaryData := SecondaryData;
  FSecondarySize := SecondarySize;
end;

destructor TJetBookmark.Destroy;
begin
  FreeMem(FData);
  inherited;
end;

function TJetBookmark.Go: Boolean;
var
  Err: JET_ERR;
begin
  if FSecondaryData = nil then
    Err := JetGotoBookmark(FTable.Database.Session.Handle, FTable.Handle, FData, FSize)
  else
    Err := JetGotoSecondaryIndexBookmark(FTable.Database.Session.Handle, FTable.Handle, FData, FSize, FSecondaryData, FSecondarySize, 0);

  if Err = 0 then
    Result := True
  else if Err = JET_errNoCurrentRecord then
    Result := False
  else
  begin
    JetError(Err);
    Result := False;
  end;
end;

end.
